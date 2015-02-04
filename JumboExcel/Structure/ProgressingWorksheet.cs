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
        public Func<Action<IEnumerable<RowLevelElement>>, IEnumerable<TProgress>> RowGenerator { get; private set; }

        const int MAX_NAME_LENGTH = 31;

        public WorksheetParametersElement Parameters { get; private set; }

        public string Name { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="name">Worksheet name. <see cref="System.String"/> of maxumum length 31.</param>
        /// <param name="parameters">Worksheet parameters.</param>
        /// <param name="rowGenerator">Generator accepting a callback, accepting rows, and returning an <see cref="IEnumerable{TProgress}"/> of progress items, which calls the row generator upon iteration.</param>
        public ProgressingWorksheet(string name, WorksheetParametersElement parameters, Func<Action<IEnumerable<RowLevelElement>>, IEnumerable<TProgress>> rowGenerator)
        {
            RowGenerator = rowGenerator;
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("name");
            if (name.Length > MAX_NAME_LENGTH)
                throw new ArgumentOutOfRangeException("name", name, "Name length must be < 32");
            Name = name;
            Parameters = parameters;
        }
    }
}