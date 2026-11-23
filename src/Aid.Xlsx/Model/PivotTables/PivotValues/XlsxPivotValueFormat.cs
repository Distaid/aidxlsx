using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.PivotTables.PivotValues;

internal class XlsxPivotValueFormat(XlsxPivotDataField pivotValue) : IXlsxPivotValueFormat
{
    private readonly XlsxPivotDataField _pivotValue = pivotValue;

    public int NumberFormatId
    {
        get
        {
            if (_pivotValue.NumberFormatValue is null)
                return -1;

            if (!XlsxPredefinedFormat.NumberFormatIds.TryGetValue(_pivotValue.NumberFormatValue.Value, out var numFmtId))
                return -1;

            return numFmtId;
        }

        set
        {
            if (!XlsxPredefinedFormat.FormatCodes.TryGetValue(value, out XlsxNumberFormat format))
                throw new ArgumentOutOfRangeException(nameof(value), value, $"Only predefined format is permitted. Use nested enums/members of {nameof(XlsxPredefinedFormat)}.");

            _pivotValue.NumberFormatValue = format;
        }
    }

    public string Format
    {
        get => _pivotValue.NumberFormatValue ?? string.Empty;
        set
        {
            _pivotValue.NumberFormatValue = XlsxNumberFormat.Parse(value);
        }
    }

    public IXlsxPivotValue SetNumberFormatId(int value)
    {
        NumberFormatId = value;
        return _pivotValue;
    }

    public IXlsxPivotValue SetFormat(string value)
    {
        Format = value;
        return _pivotValue;
    }
}
