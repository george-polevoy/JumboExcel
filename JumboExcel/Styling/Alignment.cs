using System;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Contains cell content alignment information.
    /// </summary>
    public sealed class Alignment : IEquatable<Alignment>
    {
        /// <summary>
        /// Horizontal alignment.
        /// </summary>
        public HorizontalAlignment Horizontal { get; private set; }

        /// <summary>
        /// Vertical alignment.
        /// </summary>
        public VerticalAlignment Vertical { get; private set; }

        /// <summary>
        /// Text rotation. 0-90 - as usual, 91-180 translates to the IV'ts quadrant with following formula: 90 - <see cref="TextRotation"/>.
        /// </summary>
        public int TextRotation { get; private set; }

        /// <summary>
        /// Text is wrapped by words.
        /// </summary>
        public bool WrapText { get; private set; }

        /// <summary>
        /// Creates alignment instance.
        /// </summary>
        /// <param name="horizontal">Horizontal alignment.</param>
        /// <param name="vertical">Vertical alignment.</param>
        /// <param name="textRotation">Text rotation. 0-90 - as usual, 91-180 translates to the IV'ts quadrant with following formula: 90 - <see cref="TextRotation"/>.
        /// https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.alignment(v=office.14).aspx .</param>
        /// <param name="wrapText">Text is wrapped by words.</param>
        public Alignment(HorizontalAlignment horizontal, VerticalAlignment vertical, int textRotation = 0, bool wrapText = false)
        {
            if (textRotation < 0 || textRotation > 180)
                throw new ArgumentOutOfRangeException("textRotation", textRotation, "Must be in range [0, 180] where [91,180] is transformed as 90 - textRotation");
            TextRotation = textRotation;
            WrapText = wrapText;
            Horizontal = horizontal;
            Vertical = vertical;
        }

        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <returns>
        /// true if the current object is equal to the <paramref name="other"/> parameter; otherwise, false.
        /// </returns>
        /// <param name="other">An object to compare with this object.</param>
        public bool Equals(Alignment other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Horizontal == other.Horizontal && Vertical == other.Vertical && TextRotation == other.TextRotation && WrapText == other.WrapText;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj is Alignment && Equals((Alignment)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (int)Horizontal;
                hashCode = (hashCode*397) ^ (int)Vertical;
                hashCode = (hashCode*397) ^ TextRotation;
                hashCode = (hashCode*397) ^ WrapText.GetHashCode();
                return hashCode;
            }
        }

        public static bool operator ==(Alignment left, Alignment right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(Alignment left, Alignment right)
        {
            return !Equals(left, right);
        }

        public override string ToString()
        {
            return string.Format("{0}, {1}, Rotation: {2}, WrapText: {3}", Horizontal, Vertical, TextRotation, WrapText);
        }
    }

    /// <summary>
    /// Horizontal alignment.
    /// </summary>
    public enum HorizontalAlignment
    {
        GENERAL = 0,
        LEFT = 1,
        CENTER = 2,
        RIGHT = 3,
        FILL = 4,
        JUSTIFY = 5,
        CENTER_CONTINUOUS = 6,
        DISTRIBUTED = 7
    }

    /// <summary>
    /// Vertical alignment.
    /// </summary>
    public enum VerticalAlignment
    {
        TOP = 0,
        CENTER = 1,
        BOTTOM = 2,
        JUSTIFY = 3,
        DISTRIBUTED = 4,
    }
}
