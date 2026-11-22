using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles;

internal sealed partial class XlsxBorderCellFormat
{
    private readonly XlsxCellFormat _parent;

    internal XlsxBorderCellFormat(XlsxCellFormat parent)
    {
        _parent = parent;
    }

    internal XlsxBorderStyleValues LeftBorder
    {
        get => _parent.Resolve(static x => x.Border.Left.Style);
        set => _parent.ModifyBorder(static (border, leftStyle) => border with { Left = border.Left with { Style = leftStyle } }, value);
    }

    internal XlsxColor LeftBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Left.Color);
        set => _parent.ModifyBorder(static (border, leftColor) => border with { Left = border.Left with { Color = leftColor } }, value);
    }

    internal XlsxBorderStyleValues RightBorder
    {
        get => _parent.Resolve(static x => x.Border.Right.Style);
        set => _parent.ModifyBorder(static (border, rightStyle) => border with { Right = border.Right with { Style = rightStyle } }, value);
    }

    internal XlsxColor RightBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Right.Color);
        set => _parent.ModifyBorder(static (border, rightColor) => border with { Right = border.Right with { Color = rightColor } }, value);
    }

    internal XlsxBorderStyleValues TopBorder
    {
        get => _parent.Resolve(static x => x.Border.Top.Style);
        set => _parent.ModifyBorder(static (border, topStyle) => border with { Top = border.Top with { Style = topStyle } }, value);
    }

    internal XlsxColor TopBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Top.Color);
        set => _parent.ModifyBorder(static (border, topColor) => border with { Top = border.Top with { Color = topColor } }, value);
    }

    internal XlsxBorderStyleValues BottomBorder
    {
        get => _parent.Resolve(static x => x.Border.Bottom.Style);
        set => _parent.ModifyBorder(static (border, bottomStyle) => border with { Bottom = border.Bottom with { Style = bottomStyle } }, value);
    }

    internal XlsxColor BottomBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Bottom.Color);
        set => _parent.ModifyBorder(static (border, bottomColor) => border with { Bottom = border.Bottom with { Color = bottomColor } }, value);
    }

    internal bool DiagonalUp
    {
        get => _parent.Resolve(static x => x.Border.DiagonalUp);
        set => _parent.ModifyBorder(static (border, diagonalUp) => border with { DiagonalUp = diagonalUp }, value);
    }

    internal bool DiagonalDown
    {
        get => _parent.Resolve(static x => x.Border.DiagonalDown);
        set => _parent.ModifyBorder(static (border, diagonalDown) => border with { DiagonalDown = diagonalDown }, value);
    }

    internal XlsxBorderStyleValues DiagonalBorder
    {
        get => _parent.Resolve(static x => x.Border.Diagonal.Style);
        set => _parent.ModifyBorder(static (border, diagonalStyle) => border with { Diagonal = border.Diagonal with { Style = diagonalStyle } }, value);
    }

    internal XlsxColor DiagonalBorderColor
    {
        get => _parent.Resolve(static x => x.Border.Diagonal.Color);
        set => _parent.ModifyBorder(static (border, diagonalColor) => border with { Diagonal = border.Diagonal with { Color = diagonalColor } }, value);
    }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxBorder other && (this as IEquatable<IXlsxBorder>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal void SetValue(IXlsxBorder value)
    {
        _parent.ModifyBorder((border, other) => border with
        {
            Left = new XlsxBorderLine(other.LeftBorderColor, other.LeftBorder),
            Right = new XlsxBorderLine(other.RightBorderColor, other.RightBorder),
            Top = new XlsxBorderLine(other.TopBorderColor, other.TopBorder),
            Bottom = new XlsxBorderLine(other.BottomBorderColor, other.BottomBorder),
            DiagonalUp = other.DiagonalUp,
            DiagonalDown = other.DiagonalDown,
            Diagonal = new XlsxBorderLine(other.DiagonalBorderColor, other.DiagonalBorder),
        }, value);
    }
}
