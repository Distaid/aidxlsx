using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

internal partial class XlsxAlignmentCellFormat
{
    private readonly XlsxCellFormat _parent;

    internal XlsxAlignmentCellFormat(XlsxCellFormat parent)
    {
        _parent = parent;
    }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxAlignment other && (this as IEquatable<IXlsxAlignment>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal void SetValue(IXlsxAlignment value)
    {
        Modify(static (alignment, other) => alignment with
        {
            Horizontal = other.Horizontal,
            Vertical = other.Vertical,
            TextRotation = new TextRotation(other.TextRotation),
            WrapText = other.WrapText,
            Indent = other.Indent,
            RelativeIndent = other.RelativeIndent,
            JustifyLastLine = other.JustifyLastLine,
            ShrinkToFit = other.ShrinkToFit,
            ReadingOrder = other.ReadingOrder
        }, value);
    }

    private T Resolve<T>(Func<XlsxCellFormatValue, T> selector)
    {
        return _parent.Resolve(selector);
    }

    private void Modify<TProperty>(Func<XlsxAlignmentFormatValue, TProperty, XlsxAlignmentFormatValue> modifyAlignment, TProperty value)
    {
        _parent.ModifyAlignment(modifyAlignment, value);
    }
}
