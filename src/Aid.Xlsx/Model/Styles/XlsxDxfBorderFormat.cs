using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

// TODO Styles: Plan is to have a separate interface for dxf. Implement necessary features to pass tests, then replace.
/// <summary>
/// API object for dxf hidden behind IXlsxStyle.IXlsxBorder interface.
/// </summary>
internal class XlsxDxfBorderFormat : IXlsxBorder
{
    private static readonly XlsxBorderLine NoLine = XlsxBorderLine.None;
    private readonly XlsxDxFormat _parent;

    internal XlsxDxfBorderFormat(XlsxDxFormat parent)
    {
        _parent = parent;
    }

    XlsxBorderStyleValues IXlsxBorder.OutsideBorder
    {
        set => Modify(static (border, style) => border with
        {
            Outline = true,
            Left = (border.Left ?? NoLine) with { Style = style },
            Right = (border.Right ?? NoLine) with { Style = style },
            Top = (border.Top ?? NoLine) with { Style = style },
            Bottom = (border.Bottom ?? NoLine) with { Style = style },
        }, value);
    }

    XlsxColor IXlsxBorder.OutsideBorderColor
    {
        set => Modify(static (border, color) => border with
        {
            Outline = true,
            Left = (border.Left ?? NoLine) with { Color = color },
            Right = (border.Right ?? NoLine) with { Color = color },
            Top = (border.Top ?? NoLine) with { Color = color },
            Bottom = (border.Bottom ?? NoLine) with { Color = color },
        }, value);
    }

    XlsxBorderStyleValues IXlsxBorder.InsideBorder
    {
        set => Modify(static (border, style) => border with
        {
            Outline = false,
            Left = (border.Left ?? NoLine) with { Style = style },
            Right = (border.Right ?? NoLine) with { Style = style },
            Top = (border.Top ?? NoLine) with { Style = style },
            Bottom = (border.Bottom ?? NoLine) with { Style = style },
        }, value);
    }

    XlsxColor IXlsxBorder.InsideBorderColor
    {
        set => Modify(static (border, color) => border with
        {
            Outline = false,
            Left = (border.Left ?? NoLine) with { Color = color },
            Right = (border.Right ?? NoLine) with { Color = color },
            Top = (border.Top ?? NoLine) with { Color = color },
            Bottom = (border.Bottom ?? NoLine) with { Color = color },
        }, value);
    }

    XlsxBorderStyleValues IXlsxBorder.LeftBorder
    {
        get => Resolve(static x => x.Left?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Left = (border.Left ?? NoLine) with { Style = style } }, value);
    }

    XlsxColor IXlsxBorder.LeftBorderColor
    {
        get => Resolve(static border => border.Left?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Left = (border.Left ?? NoLine) with { Color = color } }, value);
    }

    XlsxBorderStyleValues IXlsxBorder.RightBorder
    {
        get => Resolve(static x => x.Right?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Right = (border.Right ?? NoLine) with { Style = style } }, value);
    }

    XlsxColor IXlsxBorder.RightBorderColor
    {
        get => Resolve(static border => border.Right?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Right = (border.Right ?? NoLine) with { Color = color } }, value);
    }

    XlsxBorderStyleValues IXlsxBorder.TopBorder
    {
        get => Resolve(static x => x.Top?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Top = (border.Top ?? NoLine) with { Style = style } }, value);
    }

    XlsxColor IXlsxBorder.TopBorderColor
    {
        get => Resolve(static border => border.Top?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Top = (border.Top ?? NoLine) with { Color = color } }, value);
    }

    XlsxBorderStyleValues IXlsxBorder.BottomBorder
    {
        get => Resolve(static x => x.Bottom?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Bottom = (border.Bottom ?? NoLine) with { Style = style } }, value);
    }

    XlsxColor IXlsxBorder.BottomBorderColor
    {
        get => Resolve(static border => border.Bottom?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Bottom = (border.Bottom ?? NoLine) with { Color = color } }, value);
    }

    bool IXlsxBorder.DiagonalUp
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    bool IXlsxBorder.DiagonalDown
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    XlsxBorderStyleValues IXlsxBorder.DiagonalBorder
    {
        get => Resolve(static x => x.Diagonal?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Diagonal = (border.Diagonal ?? NoLine) with { Style = style } }, value);
    }

    XlsxColor IXlsxBorder.DiagonalBorderColor
    {
        get => Resolve(static border => border.Diagonal?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Diagonal = (border.Diagonal ?? NoLine) with { Color = color } }, value);
    }

    internal void SetValue(IXlsxBorder value)
    {
        Modify((border, other) => border with
        {
            Left = new XlsxBorderLine(other.LeftBorderColor, other.LeftBorder),
            Right = new XlsxBorderLine(other.RightBorderColor, other.RightBorder),
            Top = new XlsxBorderLine(other.TopBorderColor, other.TopBorder),
            Bottom = new XlsxBorderLine(other.BottomBorderColor, other.BottomBorder),
            DiagonalUp = other.DiagonalUp,
            DiagonalDown = other.DiagonalDown,
            Diagonal = new XlsxBorderLine(other.DiagonalBorderColor, other.DiagonalBorder),
            Horizontal = null,
            Vertical = null,
            Outline = true,
        }, value);
    }

    private T Resolve<T>(Func<XlsxDifferentialBorderValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Border, getProperty) ?? defaultValue;
    }

    private T Resolve<T>(Func<XlsxDifferentialBorderValue, T?> getProperty, T defaultValue)
        where T : class
    {
        return _parent.Resolve(static format => format.Border, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XlsxDifferentialBorderValue, T, XlsxDifferentialBorderValue> modify, T value)
    {
        _parent.ModifyBorder(modify, value);
    }

    IXlsxStyle IXlsxBorder.SetOutsideBorder(XlsxBorderStyleValues value)
    {
        (this as IXlsxBorder).OutsideBorder = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetOutsideBorderColor(XlsxColor value)
    {
        (this as IXlsxBorder).OutsideBorderColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetInsideBorder(XlsxBorderStyleValues value)
    {
        (this as IXlsxBorder).InsideBorder = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetInsideBorderColor(XlsxColor value)
    {
        (this as IXlsxBorder).InsideBorderColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetLeftBorder(XlsxBorderStyleValues value)
    {
        (this as IXlsxBorder).LeftBorder = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetLeftBorderColor(XlsxColor value)
    {
        (this as IXlsxBorder).LeftBorderColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetRightBorder(XlsxBorderStyleValues value)
    {
        (this as IXlsxBorder).RightBorder = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetRightBorderColor(XlsxColor value)
    {
        (this as IXlsxBorder).RightBorderColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetTopBorder(XlsxBorderStyleValues value)
    {
        (this as IXlsxBorder).TopBorder = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetTopBorderColor(XlsxColor value)
    {
        (this as IXlsxBorder).TopBorderColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetBottomBorder(XlsxBorderStyleValues value)
    {
        (this as IXlsxBorder).BottomBorder = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetBottomBorderColor(XlsxColor value)
    {
        (this as IXlsxBorder).BottomBorderColor = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetDiagonalUp()
    {
        return (this as IXlsxBorder).SetDiagonalUp(true);
    }

    IXlsxStyle IXlsxBorder.SetDiagonalUp(bool value)
    {
        (this as IXlsxBorder).DiagonalUp = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetDiagonalDown()
    {
        return (this as IXlsxBorder).SetDiagonalDown(true);
    }

    IXlsxStyle IXlsxBorder.SetDiagonalDown(bool value)
    {
        (this as IXlsxBorder).DiagonalDown = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetDiagonalBorder(XlsxBorderStyleValues value)
    {
        (this as IXlsxBorder).DiagonalBorder = value;
        return _parent;
    }

    IXlsxStyle IXlsxBorder.SetDiagonalBorderColor(XlsxColor value)
    {
        (this as IXlsxBorder).DiagonalBorderColor = value;
        return _parent;
    }

    bool IEquatable<IXlsxBorder>.Equals(IXlsxBorder? other)
    {
        throw new NotImplementedException();
    }
}
