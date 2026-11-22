using Aid.Xlsx.Styles;

namespace Aid.Xlsx.Tests.Excel.Styles;

public class FormatTestCase<TApi>
{
    private readonly Func<TApi, object> _getter;
    private readonly Action<TApi, object> _setter;
    private readonly IReadOnlyList<object> _testValues;

    private FormatTestCase(Func<TApi, object> getter, Action<TApi, object> setter, params object[] testValues)
    {
        _getter = getter;
        _setter = setter;
        _testValues = testValues;
    }

    internal static FormatTestCase<IXlsxFont> ForFont<T>(Func<IXlsxFont, T> getter, Action<IXlsxFont, T> setter, params T[] testValues)
    {
        return new FormatTestCase<IXlsxFont>(font => getter(font), (font, value) => setter(font, (T)value), testValues.Cast<object>().ToArray());
    }

    internal static FormatTestCase<IXlsxFill> ForFill<T>(Func<IXlsxFill, T> getter, Action<IXlsxFill, T> setter, params T[] testValues)
    {
        return new FormatTestCase<IXlsxFill>(fill => getter(fill), (fill, value) => setter(fill, (T)value), testValues.Cast<object>().ToArray());
    }

    internal static FormatTestCase<IXlsxBorder> ForBorder<T>(Func<IXlsxBorder, T> getter, Action<IXlsxBorder, T> setter, params T[] testValues)
    {
        return new FormatTestCase<IXlsxBorder>(border => getter(border), (border, value) => setter(border, (T)value), testValues.Cast<object>().ToArray());
    }

    internal static FormatTestCase<IXlsxAlignment> ForAlignment<T>(Func<IXlsxAlignment, T> getter, Action<IXlsxAlignment, T> setter, params T[] testValues)
    {
        return new FormatTestCase<IXlsxAlignment>(align => getter(align), (align, value) => setter(align, (T)value), testValues.Cast<object>().ToArray());
    }

    internal static FormatTestCase<IXlsxProtection> ForProtection<T>(Func<IXlsxProtection, T> getter, Action<IXlsxProtection, T> setter, params T[] testValues)
    {
        return new FormatTestCase<IXlsxProtection>(protection => getter(protection), (protection, value) => setter(protection, (T)value), testValues.Cast<object>().ToArray());
    }

    internal IEnumerable<object> Values => _testValues;

    internal object GetPropertyValue(TApi font)
    {
        return _getter(font);
    }

    internal void SetPropertyValue(TApi font, object value)
    {
        _setter(font, value);
    }
}
