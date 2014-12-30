using System;
using System.Collections.Generic;

namespace JumboExcel.Structure
{
    public class WorksheetParametersElement
    {
        public IEnumerable<ColumnElement> ColumnElements { get; private set; }

        public bool Belo { get; private set; }

        public bool Right { get; private set; }

        public WorksheetParametersElement()
        {
        }

        public WorksheetParametersElement(bool belo, bool right, IEnumerable<ColumnElement> columnElements = null)
        {
            Belo = belo;
            Right = right;
            ColumnElements = columnElements;
        }

        public WorksheetParametersElement(bool belo, bool right, params ColumnElement[] columnElements)
        {
            Belo = belo;
            Right = right;
            ColumnElements = columnElements;
        }
    }

    public class ColumnElement
    {
        public ColumnElement(int min, int max, decimal width)
        {
            if (min < 0)
                throw new ArgumentOutOfRangeException("min", min, "Must be > 0.");

            if (max < min)
                throw new ArgumentOutOfRangeException("max", max, "Must be > min.");

            if (width < 0)
                throw new ArgumentOutOfRangeException("width", width, "Must be > 0.");

            Min = min;
            Max = max;
            Width = width;
        }

        public int Min { get; private set; }
        public int Max { get; private set; }
        public decimal Width { get; set; }
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
