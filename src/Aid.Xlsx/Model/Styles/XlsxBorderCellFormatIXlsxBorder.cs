using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

internal sealed partial class XlsxBorderCellFormat : IXlsxBorder
{
    XlsxBorderStyleValues IXlsxBorder.OutsideBorder
    {
        set => _parent.ModifyOuterBorder(static (borderLine, style) => borderLine with { Style = style }, value);
    }

    XlsxColor IXlsxBorder.OutsideBorderColor
    {
        set => _parent.ModifyOuterBorder(static (borderLine, color) => borderLine with { Color = color }, value);
    }

    XlsxBorderStyleValues IXlsxBorder.InsideBorder
    {
        set => _parent.ModifyInnerBorder(static (borderLine, style) => borderLine with { Style = style }, value);
    }

    XlsxColor IXlsxBorder.InsideBorderColor
    {
        set => _parent.ModifyInnerBorder(static (borderLine, color) => borderLine with { Color = color }, value);
    }

    XlsxBorderStyleValues IXlsxBorder.LeftBorder
    {
        get => LeftBorder;
        set => LeftBorder = value;
    }

    XlsxColor IXlsxBorder.LeftBorderColor
    {
        get => LeftBorderColor;
        set => LeftBorderColor = value;
    }

    XlsxBorderStyleValues IXlsxBorder.RightBorder
    {
        get => RightBorder;
        set => RightBorder = value;
    }

    XlsxColor IXlsxBorder.RightBorderColor
    {
        get => RightBorderColor;
        set => RightBorderColor = value;
    }

    XlsxBorderStyleValues IXlsxBorder.TopBorder
    {
        get => TopBorder;
        set => TopBorder = value;
    }

    XlsxColor IXlsxBorder.TopBorderColor
    {
        get => TopBorderColor;
        set => TopBorderColor = value;
    }

    XlsxBorderStyleValues IXlsxBorder.BottomBorder
    {
        get => BottomBorder;
        set => BottomBorder = value;
    }

    XlsxColor IXlsxBorder.BottomBorderColor
    {
        get => BottomBorderColor;
        set => BottomBorderColor = value;
    }

    bool IXlsxBorder.DiagonalUp
    {
        get => DiagonalUp;
        set => DiagonalUp = value;
    }

    bool IXlsxBorder.DiagonalDown
    {
        get => DiagonalDown;
        set => DiagonalDown = value;
    }

    XlsxBorderStyleValues IXlsxBorder.DiagonalBorder
    {
        get => DiagonalBorder;
        set => DiagonalBorder = value;
    }

    XlsxColor IXlsxBorder.DiagonalBorderColor
    {
        get => DiagonalBorderColor;
        set => DiagonalBorderColor = value;
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
        // A "business" equals, when borders are visually same, they are considered equals.
        if (other is null)
            return false;

        var thisBorder = _parent.Resolve(static x => x.Border);
        var otherLeft = new XlsxBorderLine(other.LeftBorderColor, other.LeftBorder);
        if (!IsSameLine(thisBorder.Left, otherLeft))
            return false;

        var otherTop = new XlsxBorderLine(other.TopBorderColor, other.TopBorder);
        if (!IsSameLine(thisBorder.Top, otherTop))
            return false;

        var otherRight = new XlsxBorderLine(other.RightBorderColor, other.RightBorder);
        if (!IsSameLine(thisBorder.Right, otherRight))
            return false;

        var otherBottom = new XlsxBorderLine(other.BottomBorderColor, other.BottomBorder);
        if (!IsSameLine(thisBorder.Bottom, otherBottom))
            return false;

        // Check diagonals. If the diagonal flag is not set, the diagonal is not displayed. Normalize them to take into the account the direction flag
        var thisDiagonalUp = MakeThisDiagonal(thisBorder.Diagonal, thisBorder.DiagonalUp);
        var otherDiagonalUp = MakeOtherDiagonal(other, other.DiagonalUp);
        if (!IsSameLine(thisDiagonalUp, otherDiagonalUp))
            return false;

        var thisDiagonalDown = MakeThisDiagonal(thisBorder.Diagonal, thisBorder.DiagonalDown);
        var otherDiagonalDown = MakeOtherDiagonal(other, other.DiagonalDown);
        if (!IsSameLine(thisDiagonalDown, otherDiagonalDown))
            return false;

        return true;

        static bool IsSameLine(XlsxBorderLine lhs, XlsxBorderLine rhs)
        {
            if (lhs.Style == XlsxBorderStyleValues.None && rhs.Style == XlsxBorderStyleValues.None)
                return true;

            // Auto color in context of border is black, not transparent.
            return lhs.Style == rhs.Style && lhs.Color == rhs.Color;
        }

        static XlsxBorderLine MakeThisDiagonal(XlsxBorderLine diagonal, bool diagonalDirection)
        {
            return diagonal with { Style = diagonalDirection ? diagonal.Style : XlsxBorderStyleValues.None };
        }

        static XlsxBorderLine MakeOtherDiagonal(IXlsxBorder border, bool diagonalDirection)
        {
            return new XlsxBorderLine(border.DiagonalBorderColor, diagonalDirection ? border.DiagonalBorder : XlsxBorderStyleValues.None);
        }
    }
}
