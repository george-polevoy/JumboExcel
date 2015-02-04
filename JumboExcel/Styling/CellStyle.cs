using System;
using System.Drawing;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Represents style definition.
    /// </summary>
    class CellStyle : IEquatable<CellStyle>
    {
        /// <summary>
        /// Font.
        /// </summary>
        public Font Font { get; private set; }

        /// <summary>
        /// Borders.
        /// </summary>
        public Border Border { get; private set; }

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
        /// <param name="font">Font.</param>
        /// <param name="border">Borders.</param>
        /// <param name="fillColor">Fill color.</param>
        /// <param name="format">Format.</param>
        internal CellStyle(Font font, Border border, Color? fillColor, string format = null)
        {
            Font = font;
            Border = border;
            FillColor = fillColor;
            Format = format;
        }

        public bool Equals(CellStyle other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(Format, other.Format) && Equals(Font, other.Font) && Border == other.Border && FillColor.Equals(other.FillColor);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((CellStyle) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Format != null ? Format.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Font != null ? Font.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (int) Border;
                hashCode = (hashCode*397) ^ FillColor.GetHashCode();
                return hashCode;
            }
        }

        public override string ToString()
        {
            return string.Format("{0}, {1}, {2}, {3}", Font, Border, FillColor, Format);
        }
    }
}
