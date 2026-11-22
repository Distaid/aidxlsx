using Aid.Xlsx.Cells;
using static Aid.Xlsx.Styles.XlsxPredefinedFormat.DateTime;

namespace Aid.Xlsx.Styles.Formatting;

/// <summary>
/// A strongly typed number format.
/// </summary>
internal readonly record struct XlsxNumberFormat : IEquatable<string>
{
    private readonly string _format;

    internal XlsxNumberFormat(string format)
    {
        _format = format ?? throw new ArgumentNullException(nameof(format));
    }

    public string Format => _format;

    public static implicit operator string(XlsxNumberFormat value) => value._format;

    /// <summary>
    /// Is a number format a general number format?
    /// </summary>
    internal bool IsGeneralFormat()
    {
        // General format is an empty string.
        return _format.Length == 0;
    }

    internal static XlsxNumberFormat Parse(string formatCode)
    {
        // Format code was originally just a string and not checked, so keep the same semantic for now.
        return new XlsxNumberFormat(formatCode);
    }

    public bool Equals(string? other)
    {
        return other == _format;
    }

    internal XlsxDataType GetNumberDataType()
    {
        if (XlsxPredefinedFormat.NumberFormatIds.TryGetValue(this, out var formatId))
        {
            if (IsTimeOnlyFormat(formatId))
                return XlsxDataType.TimeSpan;

            if (IsDateTimeFormat(formatId))
                return XlsxDataType.DateTime;

            return XlsxDataType.Number;
        }

        if (!string.IsNullOrWhiteSpace(_format))
        {
            var dataType = GetDataTypeFromFormat();
            return dataType ?? XlsxDataType.Number;
        }

        return XlsxDataType.Number;
    }

    private static bool IsDateTimeFormat(int numberFormatId)
    {
        return (XlsxPredefinedFormat.DateTime)numberFormatId is
            DayMonthYear4WithSlashes or
            DayMonthAbbrYear2WithDashes or
            DayMonthAbbrWithDash or
            MonthDayYear4WithDashesHour24Minutes;
    }

    private static bool IsTimeOnlyFormat(int numberFormatId)
    {
        return (XlsxPredefinedFormat.DateTime)numberFormatId is
            Hour12MinutesAmPm or
            Hour12MinutesSecondsAmPm or
            Hour24Minutes or
            Hour24MinutesSeconds or
            MinutesSeconds or
            Hour12MinutesSeconds or
            MinutesSecondsMillis1;
    }

    private XlsxDataType? GetDataTypeFromFormat()
    {
        var length = _format.Length;
        var f = _format.ToLowerInvariant();
        for (var i = 0; i < length; i++)
        {
            var c = f[i];
            if (c == '"')
                i = f.IndexOf('"', i + 1);
            else if (c == '[')
            {
                // #1742 We need to skip locale prefixes in DateTime formats [...]
                i = f.IndexOf(']', i + 1);
                if (i == -1)
                    return null;
            }
            else if (c is '0' or '#' or '?')
                return XlsxDataType.Number;
            else if (c is 'y' or 'd')
                return XlsxDataType.DateTime;
            else if (c is 'h' or 's')
                return XlsxDataType.TimeSpan;
            else if (c == 'm')
            {
                // Excel treats "m" immediately after "hh" or "h" or immediately before "ss" or "s" as minutes, otherwise as a month value
                // We can ignore the "hh" or "h" prefixes as these would have been detected by the preceding condition above.
                // So we just need to make sure any 'm' is followed immediately by "ss" or "s" (excluding placeholders) to detect a timespan value
                for (var j = i + 1; j < length; j++)
                {
                    switch (f[j])
                    {
                        case 'm':
                            continue;
                        case 's':
                            return XlsxDataType.TimeSpan;
                        case >= 'a' and <= 'z':
                        case >= '0' and <= '9':
                            return XlsxDataType.DateTime;
                    }
                }
                return XlsxDataType.DateTime;
            }
        }
        return null;
    }
}
