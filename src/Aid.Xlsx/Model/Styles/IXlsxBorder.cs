#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Columns;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Rows;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Styles
{
    public enum XlsxBorderStyleValues
    {
        DashDot,
        DashDotDot,
        Dashed,
        Dotted,
        Double,
        Hair,
        Medium,
        MediumDashDot,
        MediumDashDotDot,
        MediumDashed,
        None,
        SlantDashDot,
        Thick,
        Thin
    }

    /// <summary>
    /// <para>
    /// The interface is used across many different objects. The value returned by properties is
    /// defined only for <see cref="IXlsxCell"/>, <see cref="IXlsxRow"/>, <see cref="IXlsxColumn"/>,
    /// <see cref="IXlsxWorksheet"/> and <see cref="XlsxWorkbook"/>. The returned value is undefined
    /// for other <see cref="IXlsxRangeBase"/> objects that can contain multiple different property
    /// values (e.g. <see cref="IXlsxRange"/> can contain multiple cells, <see cref="IXlsxColumns"/>
    /// can contains multiple columns).
    /// </para>
    /// </summary>
    public interface IXlsxBorder : IEquatable<IXlsxBorder>
    {
        XlsxBorderStyleValues OutsideBorder { set; }

        XlsxColor OutsideBorderColor { set; }

        XlsxBorderStyleValues InsideBorder { set; }

        XlsxColor InsideBorderColor { set; }

        /// <summary>
        /// Get or set style of the left border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XlsxBorderStyleValues.None"/>, the <see cref="LeftBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XlsxBorderStyleValues LeftBorder { get; set; }

        /// <summary>
        /// Get or set color of the left border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border is visible (=<see cref="LeftBorder"/>
        /// is not <see cref="XlsxBorderStyleValues.None"/>). Set style first, then color.
        /// </remarks>
        XlsxColor LeftBorderColor { get; set; }

        /// <summary>
        /// Get or set style of the right border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XlsxBorderStyleValues.None"/>, the <see cref="RightBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XlsxBorderStyleValues RightBorder { get; set; }

        /// <summary>
        /// Get or set color of the right border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border is visible (=<see cref="RightBorder"/>
        /// is not <see cref="XlsxBorderStyleValues.None"/>). Set style first, then color.
        /// </remarks>
        XlsxColor RightBorderColor { get; set; }

        /// <summary>
        /// Get or set style of the top border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XlsxBorderStyleValues.None"/>, the <see cref="TopBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XlsxBorderStyleValues TopBorder { get; set; }

        /// <summary>
        /// Get or set color of the top border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border is visible (=<see cref="TopBorder"/>
        /// is not <see cref="XlsxBorderStyleValues.None"/>). Set style first, then color.
        /// </remarks>
        XlsxColor TopBorderColor { get; set; }

        /// <summary>
        /// Get or set style of the bottom border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XlsxBorderStyleValues.None"/>, the <see cref="BottomBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XlsxBorderStyleValues BottomBorder { get; set; }

        /// <summary>
        /// Get or set color of the bottom border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border is visible (=<see cref="BottomBorder"/>
        /// is not <see cref="XlsxBorderStyleValues.None"/>). Set style first, then color.
        /// </remarks>
        XlsxColor BottomBorderColor { get; set; }

        Boolean DiagonalUp { get; set; }

        Boolean DiagonalDown { get; set; }

        /// <summary>
        /// Get or set style of the diagonal border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XlsxBorderStyleValues.None"/>, the <see cref="DiagonalBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XlsxBorderStyleValues DiagonalBorder { get; set; }

        /// <summary>
        /// Get or set color of the diagonal border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border line is can be visible
        /// (=<see cref="DiagonalBorder"/> is not <see cref="XlsxBorderStyleValues.None"/>).
        /// Set style first, then color.
        /// </remarks>
        XlsxColor DiagonalBorderColor { get; set; }

        IXlsxStyle SetOutsideBorder(XlsxBorderStyleValues value);

        IXlsxStyle SetOutsideBorderColor(XlsxColor value);

        IXlsxStyle SetInsideBorder(XlsxBorderStyleValues value);

        IXlsxStyle SetInsideBorderColor(XlsxColor value);

        /// <summary>
        /// Set style of the left border.
        /// </summary>
        /// <inheritdoc cref="LeftBorder"/>
        IXlsxStyle SetLeftBorder(XlsxBorderStyleValues value);

        /// <summary>
        /// Set color of the left border.
        /// </summary>
        /// <inheritdoc cref="LeftBorderColor"/>
        IXlsxStyle SetLeftBorderColor(XlsxColor value);

        /// <summary>
        /// Set style of the right border.
        /// </summary>
        /// <inheritdoc cref="RightBorder"/>
        IXlsxStyle SetRightBorder(XlsxBorderStyleValues value);

        /// <summary>
        /// Set color of the right border.
        /// </summary>
        /// <inheritdoc cref="RightBorderColor"/>
        IXlsxStyle SetRightBorderColor(XlsxColor value);

        /// <summary>
        /// Set style of the top border.
        /// </summary>
        /// <inheritdoc cref="TopBorder"/>
        IXlsxStyle SetTopBorder(XlsxBorderStyleValues value);

        /// <summary>
        /// Set color of the top border.
        /// </summary>
        /// <inheritdoc cref="TopBorderColor"/>
        IXlsxStyle SetTopBorderColor(XlsxColor value);

        /// <summary>
        /// Set style of the bottom border.
        /// </summary>
        /// <inheritdoc cref="BottomBorder"/>
        IXlsxStyle SetBottomBorder(XlsxBorderStyleValues value);

        /// <summary>
        /// Set color of the bottom border.
        /// </summary>
        /// <inheritdoc cref="BottomBorderColor"/>
        IXlsxStyle SetBottomBorderColor(XlsxColor value);

        IXlsxStyle SetDiagonalUp(); IXlsxStyle SetDiagonalUp(Boolean value);

        IXlsxStyle SetDiagonalDown(); IXlsxStyle SetDiagonalDown(Boolean value);

        /// <summary>
        /// Set style of the diagonal border.
        /// </summary>
        /// <inheritdoc cref="DiagonalBorder"/>
        IXlsxStyle SetDiagonalBorder(XlsxBorderStyleValues value);

        /// <summary>
        /// Set color of the diagonal border.
        /// </summary>
        /// <inheritdoc cref="DiagonalBorderColor"/>
        IXlsxStyle SetDiagonalBorderColor(XlsxColor value);
    }
}
