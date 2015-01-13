using System;
using System.Drawing;
using JumboExcel.Formatting;

namespace JumboExcel.Styling
{
    public sealed class NumberStyleDefinition : CellStyleDefinition
    {
        public NumberStyleDefinition(NumberFormat format, FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor) : base(fontDefinition, borderDefinition, fillColor, format.FormatCode)
        {
        }
    }

    public sealed class DateStyleDefinition : CellStyleDefinition
    {
        public DateStyleDefinition(DateTimeFormat format, FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor) : base(fontDefinition, borderDefinition, fillColor, format.FormatCode)
        {
        }
    }

    public sealed class StringStyleDefinition : CellStyleDefinition
    {
        public StringStyleDefinition(FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor) : base(fontDefinition, borderDefinition, fillColor, CommonValueFormat.String.FormatCode)
        {
        }
    }

    public sealed class SharedStringStyleDefinition : CellStyleDefinition
    {
        public SharedStringStyleDefinition(FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor)
            : base(fontDefinition, borderDefinition, fillColor, CommonValueFormat.String.FormatCode)
        {
        }
    }

    public abstract class CellStyleDefinition : IEquatable<CellStyleDefinition>
    {
        public FontDefinition FontDefinition { get; private set; }

        public BorderDefinition BorderDefinition { get; private set; }

        public Color? FillColor { get; private set; }

        internal string Format { get; private set; }

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
            return other.GetType() == GetType() && string.Equals(Format, other.Format) && Equals(FontDefinition, other.FontDefinition) && BorderDefinition == other.BorderDefinition && FillColor.Equals(other.FillColor);
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
            return string.Format("{0}, {1}, {2}, {3}, {4}", GetType().Name, FontDefinition, BorderDefinition, FillColor, Format);
        }
    }
}
