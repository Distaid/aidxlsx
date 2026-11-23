#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Sparklines;

public interface IXlsxSparklineHorizontalAxis
{
    #region Public Properties

    XlsxColor Color { get; set; }

    bool DateAxis { get; }

    bool IsVisible { get; set; }

    bool RightToLeft { get; set; }

    #endregion Public Properties

    #region Public Methods

    IXlsxSparklineHorizontalAxis SetColor(XlsxColor value);

    IXlsxSparklineHorizontalAxis SetRightToLeft(bool value);

    IXlsxSparklineHorizontalAxis SetVisible(bool value);

    #endregion Public Methods
}
