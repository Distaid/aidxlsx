#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.Sparklines
{
    internal class XlsxSparklineStyle : IXlsxSparklineStyle, IEquatable<XlsxSparklineStyle>
    {
        #region Public Properties

        public XlsxColor FirstMarkerColor { get; set; }

        public XlsxColor HighMarkerColor { get; set; }

        public XlsxColor LastMarkerColor { get; set; }

        public XlsxColor LowMarkerColor { get; set; }

        public XlsxColor MarkersColor { get; set; }

        public XlsxColor NegativeColor { get; set; }

        public XlsxColor SeriesColor { get; set; }

        #endregion Public Properties

        #region Public Methods

        public IXlsxSparklineStyle SetFirstMarkerColor(XlsxColor value)
        {
            FirstMarkerColor = value;
            return this;
        }

        public IXlsxSparklineStyle SetHighMarkerColor(XlsxColor value)
        {
            HighMarkerColor = value;
            return this;
        }

        public IXlsxSparklineStyle SetLastMarkerColor(XlsxColor value)
        {
            LastMarkerColor = value;
            return this;
        }

        public IXlsxSparklineStyle SetLowMarkerColor(XlsxColor value)
        {
            LowMarkerColor = value;
            return this;
        }

        public IXlsxSparklineStyle SetMarkersColor(XlsxColor value)
        {
            MarkersColor = value;
            return this;
        }

        public IXlsxSparklineStyle SetNegativeColor(XlsxColor value)
        {
            NegativeColor = value;
            return this;
        }

        public IXlsxSparklineStyle SetSeriesColor(XlsxColor value)
        {
            SeriesColor = value;
            return this;
        }

        #endregion Public Methods

        public static void Copy(IXlsxSparklineStyle from, IXlsxSparklineStyle to)
        {
            to.FirstMarkerColor = from.FirstMarkerColor;
            to.HighMarkerColor = from.HighMarkerColor;
            to.LastMarkerColor = from.LastMarkerColor;
            to.LowMarkerColor = from.LowMarkerColor;
            to.MarkersColor = from.MarkersColor;
            to.NegativeColor = from.NegativeColor;
            to.SeriesColor = from.SeriesColor;
        }

        #region IEquatable implementation

        /// <summary>Returns a value that indicates whether two <see cref="T:Aid.Xlsx.Sparklines.XlsxSparklineStyle" /> objects have different values.</summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>
        public static bool operator !=(XlsxSparklineStyle left, XlsxSparklineStyle right)
        {
            return !Equals(left, right);
        }

        /// <summary>Returns a value that indicates whether the values of two <see cref="T:Aid.Xlsx.Sparklines.XlsxSparklineStyle" /> objects are equal.</summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>true if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.</returns>
        public static bool operator ==(XlsxSparklineStyle left, XlsxSparklineStyle right)
        {
            return Equals(left, right);
        }

        /// <summary>Indicates whether the current object is equal to another object of the same type.</summary>
        /// <param name="other">An object to compare with this object.</param>
        /// <returns>true if the current object is equal to the <paramref name="other">other</paramref> parameter; otherwise, false.</returns>
        public bool Equals(XlsxSparklineStyle other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return FirstMarkerColor.Equals(other.FirstMarkerColor)
                   && HighMarkerColor.Equals(other.HighMarkerColor)
                   && LastMarkerColor.Equals(other.LastMarkerColor)
                   && LowMarkerColor.Equals(other.LowMarkerColor)
                   && MarkersColor.Equals(other.MarkersColor)
                   && NegativeColor.Equals(other.NegativeColor)
                   && SeriesColor.Equals(other.SeriesColor);
        }

        /// <summary>Determines whether the specified object is equal to the current object.</summary>
        /// <param name="obj">The object to compare with the current object.</param>
        /// <returns>true if the specified object  is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(XlsxSparklineStyle)) return false;
            return Equals((XlsxSparklineStyle)obj);
        }

        /// <summary>Serves as the default hash function.</summary>
        /// <returns>A hash code for the current object.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = FirstMarkerColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HighMarkerColor.GetHashCode();
                hashCode = (hashCode * 397) ^ LastMarkerColor.GetHashCode();
                hashCode = (hashCode * 397) ^ LowMarkerColor.GetHashCode();
                hashCode = (hashCode * 397) ^ MarkersColor.GetHashCode();
                hashCode = (hashCode * 397) ^ NegativeColor.GetHashCode();
                hashCode = (hashCode * 397) ^ SeriesColor.GetHashCode();
                return hashCode;
            }
        }

        #endregion IEquatable implementation
    }
}
