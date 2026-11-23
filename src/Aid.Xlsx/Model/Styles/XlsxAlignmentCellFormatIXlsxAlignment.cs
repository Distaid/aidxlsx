using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// Explicit implementation of the <see cref="IXlsxAlignment"/> interface.
/// </summary>
internal sealed partial class XlsxAlignmentCellFormat : IXlsxAlignment
{
    XlsxAlignmentHorizontalValues IXlsxAlignment.Horizontal
    {
        get => Resolve(static format => format.Alignment.Horizontal);
        set => Modify(static (alignment, hAlign) => alignment with { Horizontal = hAlign }, value);
    }

    XlsxAlignmentVerticalValues IXlsxAlignment.Vertical
    {
        get => Resolve(static format => format.Alignment.Vertical);
        set => Modify(static (alignment, vAlign) => alignment with { Vertical = vAlign }, value);
    }

    int IXlsxAlignment.Indent
    {
        get => Resolve(static format => format.Alignment.Indent);
        set => Modify(static (alignment, indent) =>
        {
            if (alignment.Horizontal == XlsxAlignmentHorizontalValues.General)
                alignment = alignment with { Horizontal = XlsxAlignmentHorizontalValues.Left };

            if (indent > 0 && !(
                               alignment.Horizontal == XlsxAlignmentHorizontalValues.Left
                            || alignment.Horizontal == XlsxAlignmentHorizontalValues.Right
                            || alignment.Horizontal == XlsxAlignmentHorizontalValues.Distributed
                        ))
            {
                throw new InvalidOperationException("For indents, only left, right, and distributed horizontal alignments are supported.");
            }

            return alignment with { Indent = indent };
        }, value);
    }

    bool IXlsxAlignment.JustifyLastLine
    {
        get => Resolve(static format => format.Alignment.JustifyLastLine);
        set => Modify(static (alignment, justifyLastLine) => alignment with { JustifyLastLine = justifyLastLine }, value);
    }

    XlsxAlignmentReadingOrderValues IXlsxAlignment.ReadingOrder
    {
        get => Resolve(static format => format.Alignment.ReadingOrder);
        set => Modify(static (alignment, readingOrder) => alignment with { ReadingOrder = readingOrder }, value);
    }

    int IXlsxAlignment.RelativeIndent
    {
        get => Resolve(static format => format.Alignment.RelativeIndent);
        set => Modify(static (alignment, relativeIndent) => alignment with { RelativeIndent = relativeIndent }, value);
    }

    bool IXlsxAlignment.ShrinkToFit
    {
        get => Resolve(static format => format.Alignment.ShrinkToFit);
        set => Modify(static (alignment, shrinkToFit) => alignment with { ShrinkToFit = shrinkToFit }, value);
    }

    int IXlsxAlignment.TextRotation
    {
        get => Resolve(static format => format.Alignment.TextRotation.Value);
        set => Modify(static (alignment, textRotation) => alignment with { TextRotation = new TextRotation(textRotation) }, value);
    }

    bool IXlsxAlignment.WrapText
    {
        get => Resolve(static format => format.Alignment.WrapText);
        set => Modify(static (alignment, wrapText) => alignment with { WrapText = wrapText }, value);
    }

    bool IXlsxAlignment.TopToBottom
    {
        get => Resolve(static format => format.Alignment.TextRotation == TextRotation.VerticalText);
        set => Modify(static (alignment, topToBottom) => alignment with { TextRotation = topToBottom ? TextRotation.VerticalText : TextRotation.None }, value);
    }

    IXlsxStyle IXlsxAlignment.SetHorizontal(XlsxAlignmentHorizontalValues value)
    {
        (this as IXlsxAlignment).Horizontal = value;
        return _parent;
    }

    IXlsxStyle IXlsxAlignment.SetVertical(XlsxAlignmentVerticalValues value)
    {
        (this as IXlsxAlignment).Vertical = value;
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

    IXlsxStyle IXlsxAlignment.SetWrapText()
    {
        return (this as IXlsxAlignment).SetWrapText(true);
    }

    IXlsxStyle IXlsxAlignment.SetWrapText(bool value)
    {
        (this as IXlsxAlignment).WrapText = value;
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

    bool IEquatable<IXlsxAlignment>.Equals(IXlsxAlignment? other)
    {
        if (other == null)
            return false;

        XlsxAlignmentFormatValue align = Resolve(static x => x.Alignment);
        if (align.Horizontal != other.Horizontal)
            return false;

        if (align.Vertical != other.Vertical)
            return false;

        if (align.Indent != other.Indent)
            return false;

        if (align.JustifyLastLine != other.JustifyLastLine)
            return false;

        if (align.ReadingOrder != other.ReadingOrder)
            return false;

        if (align.RelativeIndent != other.RelativeIndent)
            return false;

        if (align.ShrinkToFit != other.ShrinkToFit)
            return false;

        if (align.TextRotation.Value != other.TextRotation)
            return false;

        if (align.WrapText != other.WrapText)
            return false;

        return true;
    }
}
