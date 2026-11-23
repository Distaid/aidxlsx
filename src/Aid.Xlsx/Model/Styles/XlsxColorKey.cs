#nullable disable

using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Styles;

internal struct XlsxColorKey : IEquatable<XlsxColorKey>
{
    public XlsxColorType ColorType { get; set; }

    public System.Drawing.Color Color { get; set; }

    public int Indexed { get; set; }

    public XlsxThemeColor ThemeColor { get; set; }

    public double ThemeTint { get; set; }

    public override readonly int GetHashCode()
    {
        var hashCode = -331517974;
        hashCode = hashCode * -1521134295 + (int)ColorType;
        hashCode = hashCode * -1521134295 + (ColorType == XlsxColorType.Indexed ? Indexed : 0);
        hashCode = hashCode * -1521134295 + (ColorType == XlsxColorType.Theme ? (int)ThemeColor : 0);
        hashCode = hashCode * -1521134295 + (ColorType == XlsxColorType.Theme ? ThemeTint.GetHashCode() : 0);
        hashCode = hashCode * -1521134295 + (ColorType == XlsxColorType.Color ? Color.ToArgb() : 0);
        return hashCode;
    }

    public readonly bool Equals(XlsxColorKey other)
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

    public override readonly bool Equals(object obj)
    {
        if (obj is XlsxColorKey)
            return Equals((XlsxColorKey)obj);
        return base.Equals(obj);
    }

    public override readonly string ToString()
    {
        return ColorType switch
        {
            XlsxColorType.Color => Color.ToString(),
            XlsxColorType.Theme => $"{ThemeColor} ({ThemeTint})",
            XlsxColorType.Indexed => $"Indexed: {Indexed}",
            _ => base.ToString(),
        };
    }

    public static bool operator ==(XlsxColorKey left, XlsxColorKey right) => left.Equals(right);

    public static bool operator !=(XlsxColorKey left, XlsxColorKey right) => !(left.Equals(right));
}
