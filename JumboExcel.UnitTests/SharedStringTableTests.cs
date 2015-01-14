using System;
using System.Linq;
using NUnit.Framework;

namespace JumboExcel
{
    class SharedStringTableTests
    {
        [Test]
        public void CollectionAllocatesIndexes()
        {
            var collection = new SharedElementCollection<string>();
            Assert.AreEqual(0, collection.GetOrAllocateElement("Foo"));
            Assert.AreEqual(1, collection.GetOrAllocateElement("Bar"));
            Assert.AreEqual(0, collection.GetOrAllocateElement("Foo"));
        }

        [Test]
        [TestCase("a")]
        [TestCase("a, b")]
        [TestCase("a, b, c")]
        public void SharedElementCollectionCount(string source)
        {
            var expected = source.Split(new[] { ",", " " }, StringSplitOptions.RemoveEmptyEntries);
            var collection = new SharedElementCollection<string>();
            foreach (var s in expected)
                collection.GetOrAllocateElement(s);
            Assert.AreEqual(expected.Length, collection.Count);
        }

        [Test]
        [TestCase("")]
        [TestCase("apple")]
        [TestCase("apple, banana")]
        [TestCase("banana, apple")]
        public void CollectionDequeuesValuesInOrder(string source)
        {
            var expected = source.Split(new[] {",", " "}, StringSplitOptions.RemoveEmptyEntries);
            var collection = new SharedElementCollection<string>();
            foreach (var s in expected)
                collection.GetOrAllocateElement(s);

            var originalIenumerable = collection.DequeueAll();
            // ReSharper disable once PossibleMultipleEnumeration
            CollectionAssert.AreEqual(expected, originalIenumerable.ToList());
            var count = 0;
            try
            {
                // ReSharper disable once LoopCanBeConvertedToQuery
                foreach (var s in collection.DequeueAll())
                {
                    count++;
                    Console.WriteLine(s);
                }
                Assert.Fail("Should not be able to reach this point.");
            }
            catch (InvalidOperationException)
            {
                if (count > 0)
                    Assert.Fail("Should not be able to iterate at all.");
            }

            try
            {
                // ReSharper disable once PossibleMultipleEnumeration
                count = originalIenumerable.Count();
                Assert.Fail("Can't be iterated twice.");
            }
            catch (InvalidOperationException)
            {
            }

            Console.WriteLine(count);
        }
    }
}
