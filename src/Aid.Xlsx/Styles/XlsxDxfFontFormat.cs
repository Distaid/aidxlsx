using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles;

internal partial class XlsxDxfFontFormat
{
    private readonly XlsxDxFormat _parent;

    internal XlsxDxfFontFormat(XlsxDxFormat parent)
    {
        _parent = parent;
    }

    internal void SetValue(IXlsxFont value)
    {
        _parent.ModifyFont(static (font, value) => font with
        {
            Bold = value.Bold,
            Italic = value.Italic,
            Underline = value.Underline,
            Strikethrough = value.Strikethrough,
            VerticalAlignment = value.VerticalAlignment,
            Shadow = value.Shadow,
            Size = XlsxFontSize.FromPoints(value.FontSize),
            Color = value.FontColor,
            Name = value.FontName,
            Family = value.FontFamilyNumbering,
            Charset = value.FontCharSet,
            Scheme = value.FontScheme
        }, value);
    }

    private T Resolve<T>(Func<XlsxDifferentialFontValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Font, getProperty) ?? defaultValue;
    }

    private T Resolve<T>(Func<XlsxDifferentialFontValue, T?> getProperty, T defaultValue)
        where T : class
    {
        return _parent.Resolve(static format => format.Font, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XlsxDifferentialFontValue, T, XlsxDifferentialFontValue> modify, T value)
    {
        _parent.ModifyFont(modify, value);
    }
}
