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
        public IEnumerable<ColumnElement> ColumnElements { get; private set; }

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
        /// <param name="columnElements">Column configurations.</param>
        public WorksheetParametersElement(bool belo, bool right, IEnumerable<ColumnElement> columnElements = null)
        {
            Belo = belo;
            Right = right;
            ColumnElements = columnElements;
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="belo">Specifies if the summary rows are belo the grouped rows.</param>
        /// <param name="right">Specifies, if the summary columns are at the right of grouped columns (Grouped collumns are not supported in this implementation).</param>
        /// <param name="columnElements">Column configurations.</param>
        public WorksheetParametersElement(bool belo, bool right, params ColumnElement[] columnElements)
        {
            Belo = belo;
            Right = right;
            ColumnElements = columnElements;
        }
    }

    public class WorksheetElement : DocumentElement
    {
        public WorksheetParametersElement Parameters { get; private set; }
        public IEnumerable<RowLevelElement> RowsLevelElements { get; set; }
        public string Name { get; set; }
        
        public WorksheetElement(string name, WorksheetParametersElement parameters, IEnumerable<RowLevelElement> rowsLevelElements)
        {
            Name = name;
            Parameters = parameters;
            RowsLevelElements = rowsLevelElements;
        }

        public WorksheetElement(string name, WorksheetParametersElement parameters, params RowLevelElement[] rowsLevelElements)
        {
            Name = name;
            Parameters = parameters;
            RowsLevelElements = rowsLevelElements;
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
