using System;
using System.Collections.Generic;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Worksheet with progressive write capability.
    /// </summary>
    /// <typeparam name="TProgress">Progress item type.</typeparam>
    public sealed class ProgressingWorksheet<TProgress>
    {
        /// <summary>
        /// Maximim length of the worksheet name.
        /// </summary>
        const int MaxNameLength = 31;

        /// <summary>
        /// Row generation function which writes the generated row elements with provided action, and yields progress elements.
        /// </summary>
        /// <example>
        /// <code>IEnumerable&lt;int&gt; WritePortions(Action&lt;IEnumerable&lt;RowLevelElement&gt;&gt; writePortion)
        /// {
        ///     foreach (var portion in portions GetFromDatabase())
        ///     {
        ///         writePortion(portion);
        ///         yield return 0;
        ///     }
        /// }
        /// </code>
        /// </example>
        public Func<Action<IEnumerable<RowLevelElement>>, IEnumerable<TProgress>> RowGenerator { get; private set; }

        /// <summary>
        /// Worksheet parameters.
        /// </summary>
        public WorksheetParametersElement Parameters { get; private set; }

        /// <summary>
        /// Worksheet name. (1..31 characters)
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="name">Worksheet name. <see cref="System.String"/> of maxumum length 31.</param>
        /// <param name="parameters">Worksheet parameters.</param>
        /// <param name="rowGenerator">Generator accepting a callback, accepting rows, and returning an <see cref="IEnumerable{TProgress}"/> of progress items, which calls the row generator upon iteration.</param>
        public ProgressingWorksheet(string name, WorksheetParametersElement parameters, Func<Action<IEnumerable<RowLevelElement>>, IEnumerable<TProgress>> rowGenerator)
        {
            if (rowGenerator == null)
                throw new ArgumentNullException("rowGenerator");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("name");
            if (name.Length > MaxNameLength)
                throw new ArgumentOutOfRangeException("name", name, "Name length must be < 32");
            RowGenerator = rowGenerator;
            Name = name;
            Parameters = parameters;
        }
    }
}