using System;
using System.Collections.Generic;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents worksheet parameters.
    /// </summary>
    public class WorksheetParametersElement
    {
        /// <summary>
        /// Column configurations.
        /// </summary>
        public IEnumerable<ColumnConfiguration> ColumnConfigurations { get; private set; }

        /// <summary>
        /// Specifies, if the summary rows are belo the grouped rows.
        /// </summary>
        public bool Belo { get; private set; }

        /// <summary>
        /// Specifies, if the summary columns are at the right of grouped columns (Grouped collumns are not supported in this implementation).
        /// </summary>
        public bool Right { get; private set; }

        /// <summary>
        /// Default constructor.
        /// </summary>
        public WorksheetParametersElement()
        {
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="belo">Specifies if the summary rows are belo the grouped rows.</param>
        /// <param name="right">Specifies, if the summary columns are at the right of grouped columns (Grouped collumns are not supported in this implementation).</param>
        /// <param name="columnConfigurations">Column configurations.</param>
        public WorksheetParametersElement(bool belo, bool right, IEnumerable<ColumnConfiguration> columnConfigurations = null)
        {
            Belo = belo;
            Right = right;
            ColumnConfigurations = columnConfigurations;
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="belo">Specifies if the summary rows are belo the grouped rows.</param>
        /// <param name="right">Specifies, if the summary columns are at the right of grouped columns (Grouped collumns are not supported in this implementation).</param>
        /// <param name="columnConfigurations">Column configurations.</param>
        public WorksheetParametersElement(bool belo, bool right, params ColumnConfiguration[] columnConfigurations)
        {
            Belo = belo;
            Right = right;
            ColumnConfigurations = columnConfigurations;
        }
    }

    public sealed class Worksheet : DocumentElement
    {
        const int MAX_NAME_LENGTH = 31;

        public WorksheetParametersElement Parameters { get; private set; }

        public IEnumerable<RowLevelElement> RowsLevelElements { get; private set; }

        public string Name { get; private set; }
        
        public Worksheet(string name, WorksheetParametersElement parameters, IEnumerable<RowLevelElement> rowsLevelElements)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("name");
            if (name.Length > MAX_NAME_LENGTH)
                throw new ArgumentOutOfRangeException("name", name, "Name length must be < 32");
            Name = name;
            Parameters = parameters;
            RowsLevelElements = rowsLevelElements;
        }

        public Worksheet(string name, WorksheetParametersElement parameters, params RowLevelElement[] rowsLevelElements)
        {
            Name = name;
            Parameters = parameters;
            RowsLevelElements = rowsLevelElements;
        }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
