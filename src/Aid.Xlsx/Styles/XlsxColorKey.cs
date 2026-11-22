#nullable disable

using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.Styles
{
    internal struct XlsxColorKey : IEquatable<XlsxColorKey>
    {
        public XlsxColorType ColorType { get; set; }

        public System.Drawing.Color Color { get; set; }

        public int Indexed { get; set; }

        public XlsxThemeColor ThemeColor { get; set; }

        public double ThemeTint { get; set; }

        public override int GetHashCode()
        {
            var hashCode = -331517974;
            hashCode = hashCode * -1521134295 + (int)ColorType;
            hashCode = hashCode * -1521134295 + (ColorType == XlsxColorType.Indexed ? Indexed : 0);
            hashCode = hashCode * -1521134295 + (ColorType == XlsxColorType.Theme ? (int)ThemeColor : 0);
            hashCode = hashCode * -1521134295 + (ColorType == XlsxColorType.Theme ? ThemeTint.GetHashCode() : 0);
            hashCode = hashCode * -1521134295 + (ColorType == XlsxColorType.Color ? Color.ToArgb() : 0);
            return hashCode;
        }

        public bool Equals(XlsxColorKey other)
        {
            if (ColorType == other.ColorType)
            {
                if (ColorType == XlsxColorType.Color)
                {
                    // .NET Color.Equals() will return false for Color.FromArgb(255, 255, 255, 255) == Color.White
                    // Therefore we compare the ToArgb() values
                    return Color.ToArgb() == other.Color.ToArgb();
                }
                if (ColorType == XlsxColorType.Theme)
                {
                    return
                        ThemeColor == other.ThemeColor
                     && Math.Abs(ThemeTint - other.ThemeTint) < XlsxHelper.Epsilon;
                }
                return Indexed == other.Indexed;
            }

            return false;
        }

        public override bool Equals(object obj)
        {
            if (obj is XlsxColorKey)
                return Equals((XlsxColorKey)obj);
            return base.Equals(obj);
        }

        public override string ToString()
        {
            switch (ColorType)
            {
                case XlsxColorType.Color:
                    return Color.ToString();
                case XlsxColorType.Theme:
                    return $"{ThemeColor} ({ThemeTint})";
                case XlsxColorType.Indexed:
                    return $"Indexed: {Indexed}";
                default:
                    return base.ToString();
            }
        }

        public static bool operator ==(XlsxColorKey left, XlsxColorKey right) => left.Equals(right);

        public static bool operator !=(XlsxColorKey left, XlsxColorKey right) => !(left.Equals(right));
    }
}
