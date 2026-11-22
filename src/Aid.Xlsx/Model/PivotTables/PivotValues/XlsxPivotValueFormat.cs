using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.PivotTables.PivotValues;

internal class XlsxPivotValueFormat : IXlsxPivotValueFormat
{
    private readonly XlsxPivotDataField _pivotValue;

    public XlsxPivotValueFormat(XlsxPivotDataField pivotValue)
    {
        _pivotValue = pivotValue;
    }

    public Int32 NumberFormatId
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
            if (!XlsxPredefinedFormat.FormatCodes.TryGetValue(value, out var format))
                throw new ArgumentOutOfRangeException($"Only predefined format is permitted. Use nested enums/members of {nameof(XlsxPredefinedFormat)}.");

            _pivotValue.NumberFormatValue = format;
        }
    }

    public String Format
    {
        get => _pivotValue.NumberFormatValue ?? string.Empty;
        set
        {
            _pivotValue.NumberFormatValue = XlsxNumberFormat.Parse(value);
        }
    }

    public IXlsxPivotValue SetNumberFormatId(Int32 value)
    {
        NumberFormatId = value;
        return _pivotValue;
    }

    public IXlsxPivotValue SetFormat(String value)
    {
        Format = value;
        return _pivotValue;
    }
}
