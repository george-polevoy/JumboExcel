using System;
using System.Drawing;

namespace JumboExcel.Styling
{
    public sealed class FontDefinition : IEquatable<FontDefinition>
    {
        public string Typeface { get; private set; }

        public decimal Size { get; private set; }

        public Color Color { get; private set; }

        public FontWeight FontWeight { get; private set; }

        public FontSlope FontSlope { get; private set; }

        public FontDefinition(string typeface, decimal size, Color color, FontSlope fontSlope, FontWeight fontWeight)
        {
            if (size < 0.1m || size > 500m)
                throw new ArgumentOutOfRangeException("size", size, "Must be in range (0.1, 500.0)");
            Typeface = typeface;
            Size = size;
            Color = color;
            FontWeight = fontWeight;
            FontSlope = fontSlope;
        }

        public bool Equals(FontDefinition other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(Typeface, other.Typeface) && Size == other.Size && Color.Equals(other.Color) && FontWeight == other.FontWeight && FontSlope == other.FontSlope;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj is FontDefinition && Equals((FontDefinition) obj);
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

        public static bool operator ==(FontDefinition left, FontDefinition right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(FontDefinition left, FontDefinition right)
        {
            return !Equals(left, right);
        }

        public override string ToString()
        {
            return string.Format("{0}, {1}, {2}, {3}, {4}", Typeface, Size, Color, FontWeight, FontSlope);
        }
    }
}
