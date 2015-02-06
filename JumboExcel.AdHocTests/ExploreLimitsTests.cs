using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using JumboExcel.Structure;
using JumboExcel.Styling;
using NUnit.Framework;

namespace JumboExcel
{
    /// <summary>
    /// Explores performance for various excel sheet widths and heights. Reports the results in a beautiful logarithmic scale diagram.
    /// </summary>
    class ExploreLimitsTests
    {
        // Can try smaller constance for debugging.
        //private const int Timeout = 100;
        //private const int MaxCells = 200000;
        //private const int MinCells = 80;
        //private const int MaxRows = 8000;
        //private const int MaxColumns = 8000;
        //private const int ScaleDivisor = 10;

        /// <summary>
        /// Limit the execution time. We are interested only in times belo 2000. Some results will exceed this limit, but the exploration would stop at that point.
        /// </summary>
        private const int Timeout = 2000;

        /// <summary>
        /// We are aiming at generating a meg of cells in specified timeout.
        /// </summary>
        private const int MaxCells = 1000000;
        /// <summary>
        /// Not interested in the results belo.
        /// </summary>
        private const int MinCells = 1000;

        private const int MaxRows = 1000000;
        private const int MaxColumns = 1000000;

        /// <summary>
        /// Header style for performance diagram.
        /// </summary>
        readonly NumberStyle perfHeaderStyle = new NumberStyle(null, null, Border.NONE, Color.Wheat);

        /// <summary>
        /// Visually exploring performance. Exercises various worksheet widths and heights.
        /// Runs for a long time, reporting it's progress in console output, than generates a nice Excel diagram and uses shell execute to open it in Excel.
        /// </summary>
        /// <remarks>
        /// Digram would be a very large hyperbolic shape, so using logarithmic scale for width and height.
        /// </remarks>
        [Test, Explicit]
        public void ExploreWorksheetSizeGenerationPerformance()
        {
            var timings = new Dictionary<Tuple<int,int>, long>();
            foreach (var h in ExploreUpperBound(h=> CanGenerateHeight(h, timings), 1))
            {
                Console.WriteLine("Height {0} is ok.", h);
            }

            var timingsSummary =
                from timing in timings
                group timing by Tuple.Create(Scale(timing.Key.Item1), Scale(timing.Key.Item2)) into g
                select new {cell = g.Key, average = g.Average(i=>i.Value)};

            var timingsDictionary = timingsSummary.ToDictionary(i => i.cell, i => i.average);

            var rowHeaders =
                from timing in timings
                group timing by Scale(timing.Key.Item1)
                into g
                select new {logIndex = g.Key, avgRow = g.Average(i=>i.Key.Item1)};

            var rowHeadersDictionary = rowHeaders.ToDictionary(i => i.logIndex, i => i.avgRow);

            var columnHeaders =
                from timing in timings
                group timing by Scale(timing.Key.Item2)
                into g
                select new {logIndex = g.Key, avgColumn = g.Average(i=>i.Key.Item2)};

            var columnHeadersDictionary = columnHeaders.ToDictionary(i => i.logIndex, i => i.avgColumn);

            TestHelper.WriteAndExecuteExcel(new[]
            {
                new Worksheet("Timings", new WorksheetParametersElement(),
                    GetPerfSummaryRows(timingsDictionary, rowHeadersDictionary, columnHeadersDictionary))
            });
        }

        static NumberStyle CreateNumberStyle(double intensity)
        {
            var component = (int)Math.Max(Math.Min((1-intensity) * 255, 255), 0);
            return new NumberStyle(null, null, Border.NONE, Color.FromArgb(128 + component / 2, 128 + component / 2, 255));
        }

        CellElement CreatePerfHeaderCell(double? value, NumberStyle style)
        {
            return new DecimalCell(value.HasValue ? (decimal?)value.Value : null, style);
        }

        private IEnumerable<Row> GetPerfSummaryRows(Dictionary<Tuple<int, int>, double> timingsDictionary, Dictionary<int, double> rowHeadersDictionary, Dictionary<int, double> columnHeadersDictionary)
        {
            var cellElements = new List<CellElement>{EmptyCell.Instance};
            cellElements.AddRange(Enumerable.Range(1, Scale(MaxColumns)).Select(i => CreatePerfHeaderCell(TryGet(rowHeadersDictionary, i), perfHeaderStyle)));
            yield return new Row(cellElements);

            foreach (var y in Enumerable.Range(1, Scale(MaxRows)))
            {
                yield return CreatePerfRow(timingsDictionary, y, columnHeadersDictionary);
            }
        }

        private Row CreatePerfRow(Dictionary<Tuple<int, int>, double> timingsDictionary, int y, Dictionary<int, double> columnHeadersDictionary)
        {
            var cellElements = new List<CellElement> { CreatePerfHeaderCell(TryGet(columnHeadersDictionary, y), perfHeaderStyle) };
            cellElements.AddRange(Enumerable.Range(1, Scale(MaxColumns)).Select(x => Tuple.Create(x, y)).Select(cell => timingsDictionary.ContainsKey(cell) ? (CellElement)CreatePerfValueCell((decimal)timingsDictionary[cell]) : EmptyCell.Instance));
            return new Row(cellElements);
        }

        private static DecimalCell CreatePerfValueCell(decimal number)
        {
            return new DecimalCell(number, CreateNumberStyle((double)number / Timeout));
        }

        int Scale(long dimension)
        {
            return (int) Math.Log(dimension, 2);
            //return (int)(dimension / ScaleDivisor);
        }

        private bool CanGenerateHeight(long h, Dictionary<Tuple<int,int>, long> timings)
        {
            var exploreUpperBound = ExploreUpperBound(w => CanGenerateWidth(w, h, timings), 1).ToList();
            return exploreUpperBound.Count > 0;
        }

        private bool CanGenerateWidth(long w, long h, Dictionary<Tuple<int, int>, long> timings)
        {
            if (w*h < MinCells)
                return true;
            if (w*h > MaxCells)
                return false;

            var rows = h;
            var columns = w;

            if (rows > MaxRows)
                return false;
            if (columns > MaxColumns)
                return false;

            var sw = Stopwatch.StartNew();

            using (var memoryStream = new MemoryStream(4096))
            {
                OpenXmlBuilder.Write(memoryStream, new[]
                {
                    new Worksheet(
                        "Huge",
                        new WorksheetParametersElement(),
                        Enumerable.Range(0, (int)rows).Select(row => new Row(
                            Enumerable.Range(0, (int)columns).Select(column => new IntegerCell(GenerateNumber(row, column))))))
                });
                sw.Stop();
                if (sw.ElapsedMilliseconds > Timeout * 0.7)
                {
                    Console.WriteLine("Rows: {0}, Columns: {1}, File size: {2}", rows, columns, memoryStream.Length);
                }
            }
            
            timings[Tuple.Create((int)rows, (int)columns)] = sw.ElapsedMilliseconds;
                
            return sw.ElapsedMilliseconds <= Timeout;
        }

        [Test, Explicit]
        [TestCase(0, 0)]
        [TestCase(0, 1)]
        [TestCase(1, 1)]
        [TestCase(0, 2)]
        [TestCase(1, 2)]
        [TestCase(0, 3)]
        [TestCase(0, 10)]
        [TestCase(0, 63)]
        [TestCase(1, 63)]
        [TestCase(2, 63)]
        [TestCase(0, 64)]
        [TestCase(0, 65)]
        [TestCase(0, 100)]
        [TestCase(0, 1000000)]
        public void ExploreUpperBoundTest(long first, long last)
        {
            var possibleValuesFound = ExploreUpperBound(i=>i<=last, first).ToList();
            var allPossibleValues = new HashSet<long>(Enumerable.Range((int) first, (int) (last - first + 1)).Select(i => (long) i));
            CollectionAssert.AllItemsAreUnique(possibleValuesFound);
            CollectionAssert.IsSubsetOf(possibleValuesFound, allPossibleValues);
            CollectionAssert.AreEqual(possibleValuesFound.OrderBy(i => i), possibleValuesFound);
            Assert.AreEqual(last, possibleValuesFound.Last());
            Assert.LessOrEqual(possibleValuesFound.Count, 2 * Math.Log(last + 1, 2) + 1);
        }

        /// <summary>
        /// Explores the boolean function to find the upper index, at which it returns false.
        /// </summary>
        /// <param name="check">Boolean function to explore.</param>
        /// <param name="first">Lower index.</param>
        /// <returns>Returns a sequence of successful explored values in ascending order.</returns>
        static IEnumerable<long> ExploreUpperBound(Func<long, bool> check, long first)
        {
            if (!check(first))
                yield break;
            yield return first;
            var current = 1;
            while(current + first < long.MaxValue / 2)
            {
                if (check(current + first))
                    yield return current + first;
                else
                {
                    current /= 2;
                    var convergeStep = current / 2;
                    while (convergeStep > 0)
                    {
                        if (check(current + convergeStep + first))
                        {
                            yield return current + convergeStep + first;
                            current += convergeStep;
                        }
                        convergeStep /= 2;
                    }
                    yield break;
                }
                current *= 2;
            }
        }

        static long GenerateNumber(int row, int column)
        {
            return row < column ? (long)row * column : - (long)row * column;
        }

        [Test, Explicit]
        public void ExploreTestSizes()
        {
            foreach (var testSize in GetTestSizes())
            {
                Console.WriteLine(testSize);
            }
        }

        static IEnumerable<int> GetTestSizes()
        {
            yield return 1;
            for (var i = 1; i < 21; i++)
            {
                yield return 1 << i;
                yield return (1 << i) + (1 << (i - 1));
            }
        }

        static T? TryGet<T>(IDictionary<int, T> d, int index) where T : struct
        {
            T element;
            if (d.TryGetValue(index, out element))
                return element;
            return null;
        }
    }
}