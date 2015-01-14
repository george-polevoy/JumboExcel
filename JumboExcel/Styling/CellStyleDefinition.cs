using System;
using System.Drawing;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Represents style definition.
    /// </summary>
    class CellStyleDefinition : IEquatable<CellStyleDefinition>
    {
        /// <summary>
        /// Font.
        /// </summary>
        public FontDefinition FontDefinition { get; private set; }

        /// <summary>
        /// Borders.
        /// </summary>
        public BorderDefinition BorderDefinition { get; private set; }

        /// <summary>
        /// Fill color.
        /// </summary>
        public Color? FillColor { get; private set; }

        /// <summary>
        /// Format.
        /// </summary>
        internal string Format { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="fontDefinition">Font.</param>
        /// <param name="borderDefinition">Borders.</param>
        /// <param name="fillColor">Fill color.</param>
        /// <param name="format">Format.</param>
        internal CellStyleDefinition(FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor, string format = null)
        {
            FontDefinition = fontDefinition;
            BorderDefinition = borderDefinition;
            FillColor = fillColor;
            Format = format;
        }

        public bool Equals(CellStyleDefinition other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(Format, other.Format) && Equals(FontDefinition, other.FontDefinition) && BorderDefinition == other.BorderDefinition && FillColor.Equals(other.FillColor);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((CellStyleDefinition) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Format != null ? Format.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (FontDefinition != null ? FontDefinition.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (int) BorderDefinition;
                hashCode = (hashCode*397) ^ FillColor.GetHashCode();
                return hashCode;
            }
        }

        public override string ToString()
        {
            return string.Format("{0}, {1}, {2}, {3}", FontDefinition, BorderDefinition, FillColor, Format);
        }
    }
}
