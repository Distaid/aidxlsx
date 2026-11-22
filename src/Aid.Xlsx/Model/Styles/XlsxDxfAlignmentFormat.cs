using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles;

internal class XlsxDxfAlignmentFormat : IXlsxAlignment
{
    private readonly XlsxAlignmentFormatValue _default = XlsxAlignmentFormatValue.Default;
    private readonly XlsxDxFormat _parent;

    internal XlsxDxfAlignmentFormat(XlsxDxFormat parent)
    {
        _parent = parent;
    }

    XlsxAlignmentHorizontalValues IXlsxAlignment.Horizontal
    {
        get => Resolve(static alignment => alignment.Horizontal, _default.Horizontal);
        set => Modify(static (alignment, horizontal) => alignment with { Horizontal = horizontal }, value);
    }

    XlsxAlignmentVerticalValues IXlsxAlignment.Vertical
    {
        get => Resolve(static alignment => alignment.Vertical, _default.Vertical);
        set => Modify(static (alignment, vertical) => alignment with { Vertical = vertical }, value);
    }

    int IXlsxAlignment.Indent
    {
        get => Resolve(static alignment => alignment.Indent, _default.Indent);
        set => Modify(static (alignment, indent) => alignment with { Indent = indent }, value);
    }

    bool IXlsxAlignment.JustifyLastLine
    {
        get => Resolve(static alignment => alignment.JustifyLastLine, _default.JustifyLastLine);
        set => Modify(static (alignment, justifyLastLine) => alignment with { JustifyLastLine = justifyLastLine }, value);
    }

    XlsxAlignmentReadingOrderValues IXlsxAlignment.ReadingOrder
    {
        get => Resolve(static alignment => alignment.ReadingOrder, _default.ReadingOrder);
        set => Modify(static (alignment, readingOrder) => alignment with { ReadingOrder = readingOrder }, value);
    }

    int IXlsxAlignment.RelativeIndent
    {
        get => Resolve(static alignment => alignment.RelativeIndent, _default.RelativeIndent);
        set => Modify(static (alignment, relativeIndent) => alignment with { RelativeIndent = relativeIndent }, value);
    }

    bool IXlsxAlignment.ShrinkToFit
    {
        get => Resolve(static alignment => alignment.ShrinkToFit, _default.ShrinkToFit);
        set => Modify(static (alignment, shrinkToFit) => alignment with { ShrinkToFit = shrinkToFit }, value);
    }

    int IXlsxAlignment.TextRotation
    {
        get => Resolve(static alignment => alignment.TextRotation?.Value, _default.TextRotation.Value);
        set => Modify(static (alignment, textRotation) => alignment with { TextRotation = new TextRotation(textRotation) }, value);
    }

    bool IXlsxAlignment.WrapText
    {
        get => Resolve(static alignment => alignment.WrapText, _default.WrapText);
        set => Modify(static (alignment, wrapText) => alignment with { WrapText = wrapText }, value);
    }

    bool IXlsxAlignment.TopToBottom
    {
        get => Resolve(static alignment => alignment.TextRotation == TextRotation.VerticalText, false);
        set => Modify(static (alignment, topToBottom) => alignment with { TextRotation = topToBottom ? TextRotation.VerticalText : TextRotation.None }, value);
    }

    bool IEquatable<IXlsxAlignment>.Equals(IXlsxAlignment? other)
    {
        throw new NotSupportedException();
    }

    IXlsxStyle IXlsxAlignment.SetHorizontal(XlsxAlignmentHorizontalValues value)
    {
        (this as IXlsxAlignment).Horizontal = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetIndent(int value)
    {
        (this as IXlsxAlignment).Indent = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetJustifyLastLine()
    {
        return (this as IXlsxAlignment).SetJustifyLastLine(true);
    }

    IXlsxStyle IXlsxAlignment.SetJustifyLastLine(bool value)
    {
        (this as IXlsxAlignment).JustifyLastLine = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetReadingOrder(XlsxAlignmentReadingOrderValues value)
    {
        (this as IXlsxAlignment).ReadingOrder = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetRelativeIndent(int value)
    {
        (this as IXlsxAlignment).RelativeIndent = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetShrinkToFit()
    {
        return (this as IXlsxAlignment).SetShrinkToFit(true);
    }

    IXlsxStyle IXlsxAlignment.SetShrinkToFit(bool value)
    {
        (this as IXlsxAlignment).ShrinkToFit = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetTextRotation(int value)
    {
        (this as IXlsxAlignment).TextRotation = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetTopToBottom()
    {
        return (this as IXlsxAlignment).SetTopToBottom(true);
    }

    IXlsxStyle IXlsxAlignment.SetTopToBottom(bool value)
    {
        (this as IXlsxAlignment).TopToBottom = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetVertical(XlsxAlignmentVerticalValues value)
    {
        (this as IXlsxAlignment).Vertical = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetWrapText()
    {
        return (this as IXlsxAlignment).SetWrapText(true);
    }

    IXlsxStyle IXlsxAlignment.SetWrapText(bool value)
    {
        (this as IXlsxAlignment).WrapText = value;
        return _parent;
    }

    internal void SetValue(IXlsxAlignment value)
    {
        _parent.ModifyAlignment(static (alignment, value) => alignment with
        {
            Horizontal = value.Horizontal,
            Vertical = value.Vertical,
            TextRotation = new TextRotation(value.TextRotation),
            WrapText = value.WrapText,
            Indent = value.Indent,
            RelativeIndent = value.RelativeIndent,
            JustifyLastLine = value.JustifyLastLine,
            ShrinkToFit = value.ShrinkToFit,
            ReadingOrder = value.ReadingOrder,
        }, value);
    }

    private T Resolve<T>(Func<XlsxDifferentialAlignmentValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Alignment, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XlsxDifferentialAlignmentValue, T, XlsxDifferentialAlignmentValue> modify, T value)
    {
        _parent.ModifyAlignment(modify, value);
    }
}
