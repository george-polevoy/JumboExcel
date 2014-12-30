using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting;
using JumboExcel.Styling;
using NUnit.Framework;

namespace JumboExcel
{
    class CellStyleTests
    {
        [Test]
        public void InequalityOfDifferentTypes()
        {
            var a =
                new StringStyleDefinition(
                    new FontDefinition("arial", 11, Color.White, FontSlope.Normal, FontWeight.Normal), BorderDefinition.None,
                    null);
            var b = new NumberStyleDefinition(null,
                new FontDefinition("arial", 11, Color.White, FontSlope.Normal, FontWeight.Normal), BorderDefinition.None,
                null);

            Assert.AreNotEqual(a,b);
        }

        [Test]
        public void TestCellStyleIdentity()
        {
            var q =
                from typeface in new[] {null, "Arial", "Calibri"}
                from size in new[] {11m, 16m}
                from color in new[] {Color.Black, Color.White}
                from fontSlope in new[] {FontSlope.Normal, FontSlope.Italic}
                from fontWeight in new[] {FontWeight.Normal, FontWeight.Bold}
                from border in new[] {BorderDefinition.None, BorderDefinition.All}
                from fillColor in new Color?[] {null, Color.Bisque, Color.Azure}
                from definitions in new CellStyleDefinition[]
                {
                    new StringStyleDefinition(new FontDefinition(typeface, size, color, fontSlope, fontWeight), border, fillColor),
                    new NumberStyleDefinition("1", new FontDefinition(typeface, size, color, fontSlope, fontWeight), border, fillColor),
                    new NumberStyleDefinition("0", new FontDefinition(typeface, size, color, fontSlope, fontWeight), border, fillColor),
                    new NumberStyleDefinition(null, new FontDefinition(typeface, size, color, fontSlope, fontWeight), border, fillColor),
                    new DateStyleDefinition("yyyy", new FontDefinition(typeface, size, color, fontSlope, fontWeight), border, fillColor),
                    new DateStyleDefinition(null, new FontDefinition(typeface, size, color, fontSlope, fontWeight), border, fillColor),
                }
                select definitions;

            var items = q.ToArray();

            Console.WriteLine(items.Length);

            foreach (var left in items)
            {
                Assert.IsFalse(left.Equals(null));
                foreach (var right in items)
                {
                    if (ReferenceEquals(left, right))
                    {
                        Assert.IsTrue(left.Equals(right));
                        Assert.IsTrue(right.Equals(left));
                        Assert.AreEqual(left.GetHashCode(), right.GetHashCode());
                    }
                    else
                    {
                        Assert.AreNotEqual(left, right);
                    }
                }
            }

            var distinctHashCodes = items.Select(i => i.GetHashCode()).Distinct().Count();
            Console.WriteLine("Inequal items count: {0}, Distinct hash codes: {1}", items.Length, distinctHashCodes);
            Assert.IsTrue(distinctHashCodes > items.Length * 0.6);
        }
    }
}
