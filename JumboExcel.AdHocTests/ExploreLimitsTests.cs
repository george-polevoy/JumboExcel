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
    class ExploreLimitsTests
    {
        private const int Timeout = 1000;
        private const int MaxCells = 500000;
        private const int MinCells = 4000;
        private const int MaxRows = 4000;
        private const int MaxColumns = 4000;
        private const int ScaleDivisor = 100;

        [Test, Explicit]
        public void ExploreFileSize()
        {
            var timings = new Dictionary<Tuple<int,int>, long>();
            foreach (var h in ExploreUpperBound(h=> CanGenerateHeight(h + 1, timings)))
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
                    GetSummaryRows(timingsDictionary, rowHeadersDictionary, columnHeadersDictionary))
            });
        }

        T? GetFromDictionary<T>(Dictionary<int, T> d, int index) where T : struct
        {
            T element;
            if (d.TryGetValue(index, out element))
                return element;
            return null;
        }

        readonly NumberStyle headerStyle = new NumberStyle(null, null, Border.NONE, Color.Wheat);

        static NumberStyle CreateNumberStyle(double intensity)
        {
            var component = (int)Math.Max(Math.Min((1-intensity) * 255, 255), 0);
            return new NumberStyle(null, null, Border.NONE, Color.FromArgb(128 + component / 2, 128 + component / 2, 255));
        }

        CellElement CreateCell(double? value, NumberStyle style)
        {
            return new DecimalCell(value.HasValue ? (decimal?)value.Value : null, style);
        }

        private IEnumerable<Row> GetSummaryRows(Dictionary<Tuple<int, int>, double> timingsDictionary, Dictionary<int, double> rowHeadersDictionary, Dictionary<int, double> columnHeadersDictionary)
        {
            var cellElements = new List<CellElement>{EmptyCell.Instance};
            cellElements.AddRange(Enumerable.Range(1, Scale(MaxColumns)).Select(i => CreateCell(GetFromDictionary(rowHeadersDictionary, i), headerStyle)));
            yield return new Row(cellElements);

            foreach (var y in Enumerable.Range(1, Scale(MaxRows)))
            {
                yield return CreateRow(timingsDictionary, y, columnHeadersDictionary);
            }
        }

        private Row CreateRow(Dictionary<Tuple<int, int>, double> timingsDictionary, int y, Dictionary<int, double> columnHeadersDictionary)
        {
            var cellElements = new List<CellElement> { CreateCell(GetFromDictionary(columnHeadersDictionary, y), headerStyle) };
            cellElements.AddRange(Enumerable.Range(1, Scale(MaxColumns)).Select(x => Tuple.Create(x, y)).Select(cell => timingsDictionary.ContainsKey(cell) ? (CellElement)CreateValueCell((decimal)timingsDictionary[cell]) : EmptyCell.Instance));
            return new Row(cellElements);
        }

        private static DecimalCell CreateValueCell(decimal number)
        {
            return new DecimalCell(number, CreateNumberStyle((double)number / Timeout));
        }

        int Scale(long dimension)
        {
            //return (int) Math.Log(dimension, 2);
            return (int)(dimension / ScaleDivisor);
        }

        private bool CanGenerateHeight(long h, Dictionary<Tuple<int,int>, long> timings)
        {
            var exploreUpperBound = ExploreUpperBound(w => CanGenerateWidth(w + 1, h, timings)).ToList();
            return exploreUpperBound.Count > 1;
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
        [TestCase(0)]
        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [TestCase(10)]
        [TestCase(63)]
        [TestCase(64)]
        [TestCase(65)]
        [TestCase(100)]
        [TestCase(1000000)]
        public void ExploreUpperBound(long bound)
        {
            var exploreUpperBound = ExploreUpperBound(i=>i<=bound).ToList();
            Assert.AreEqual(bound, exploreUpperBound.Last());
            Assert.LessOrEqual(exploreUpperBound.Count, 2 * Math.Log(bound + 1, 2) + 1);
        }

        static IEnumerable<long> ExploreUpperBound(Func<long, bool> check)
        {
            if (!check(0))
                yield break;
            yield return 0;
            var current = 1;
            while(current < long.MaxValue)
            {
                if (check(current))
                    yield return current;
                else
                {
                    current /= 2;
                    var convergeStep = current / 2;
                    while (convergeStep > 0)
                    {
                        if (check(current + convergeStep))
                        {
                            yield return current + convergeStep;
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
    }
}