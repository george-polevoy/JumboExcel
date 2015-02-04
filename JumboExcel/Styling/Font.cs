using System;
using System.Drawing;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Represents font definition.
    /// </summary>
    public sealed class Font : IEquatable<Font>
    {
        /// <summary>
        /// Typeface of the font.
        /// </summary>
        public string Typeface { get; private set; }

        /// <summary>
        /// Size of the font in default OpenXml size units.
        /// </summary>
        public decimal Size { get; private set; }

        /// <summary>
        /// Color of the text in a cell.
        /// </summary>
        public Color Color { get; private set; }

        /// <summary>
        /// Font weight.
        /// </summary>
        public FontWeight FontWeight { get; private set; }

        /// <summary>
        /// Font slope.
        /// </summary>
        public FontSlope FontSlope { get; private set; }

        /// <summary>
        /// Constructs font definition.
        /// </summary>
        /// <param name="typeface">Typeface.</param>
        /// <param name="size">Size.</param>
        /// <param name="color">Color.</param>
        /// <param name="fontSlope">Slope.</param>
        /// <param name="fontWeight">Weight.</param>
        public Font(string typeface, decimal size, Color color, FontSlope fontSlope, FontWeight fontWeight)
        {
            if (size < 0.1m || size > 500m)
                throw new ArgumentOutOfRangeException("size", size, "Must be in range (0.1, 500.0)");
            Typeface = typeface;
            Size = size;
            Color = color;
            FontWeight = fontWeight;
            FontSlope = fontSlope;
        }

        public bool Equals(Font other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(Typeface, other.Typeface) && Size == other.Size && Color.Equals(other.Color) && FontWeight == other.FontWeight && FontSlope == other.FontSlope;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj is Font && Equals((Font) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Typeface != null ? Typeface.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ Size.GetHashCode();
                hashCode = (hashCode*397) ^ Color.GetHashCode();
                hashCode = (hashCode*397) ^ (int) FontWeight;
                hashCode = (hashCode*397) ^ (int) FontSlope;
                return hashCode;
            }
        }

        public static bool operator ==(Font left, Font right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(Font left, Font right)
        {
            return !Equals(left, right);
        }

        public override string ToString()
        {
            return string.Format("{0}, {1}, {2}, {3}, {4}", Typeface, Size, Color, FontWeight, FontSlope);
        }
    }
}
